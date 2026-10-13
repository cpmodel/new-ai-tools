"""Build CPAT_Industry_Kernel_Egypt_v0.9.xlsx from v0.8 (Task D: IPCC-inferred process-emission semi-elasticities).

Manual inputs rows 49-61 (new section): one parameter set per CBAM product and process category (np = non-fuel
process CO2, no = non-CO2 process), each documented with its abatement lever, source and confidence:

  ER(P) = ERmax x (1 - exp(-beta x P))        share of covered process emissions abated at carbon price P ($/tCO2e)
  beta  = -ln(1 - ER*/ERmax) / P*             calibrated to an anchor: ER* abated at price P*
  semi-elasticity at P = 0 = ERmax x beta     (% change in emissions per $1, the user's definition)

ERmax (technical potential by ~2035) caps abatement, so high prices (RAMP_EG3 reaches 102.5 in 2041) cannot abate
more than the available technology. ERmax = 1 gives a constant semi-elasticity beta.

Selector Manual inputs E50: IPCC (default) or ADHOC. ADHOC uses the AdHocCalculations values in rows 40-47 (one beta
for np + no, ERmax = 1) and reproduces v0.8 exactly. Effective parameters: Manual inputs T:W.

v0.9 ships PLACEHOLDER anchors (ERmax = 1, P* = 100, ER* = ad hoc ER at $100; review fill), so IPCC = ADHOC and all
results equal v0.8. Values from the separate IPCC research report go into Manual inputs E:G and I:K, rows 53-60.

CBAM block (both scenarios), rows o118-o125: E..H link to the effective beta.np, ERmax.np, beta.no, ERmax.no;
process ER = -prod/1000 x sum over np, no of [c.cat x max(EF.cat, 0) x ERmax.cat x (1 - exp(-beta.cat x P))], where
P = process price seen by the product (o65 flags x pptraj o63 + ETS). Identical to v0.8 when ADHOC is selected.
The fund placeholder in o107-o114 (fuel ER at the fund shadow price) still uses rows 40-47 (Task G/J).

Run with Excel installed:  python build_v0_9.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, DATA_COLS, REVIEW, TAN, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.8.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.8.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.9.xlsx")

PRODUCTS = ("DRI-EAF steel", "Scrap-EAF steel", "BF-BOF steel (reference)", "Grey clinker (dry-process)",
            "Ammonia (net merchant)", "Urea", "Ammonium nitrate", "Primary aluminium")

# (np: ERmax, P*, ER*, lever/source), (no: ERmax, P*, ER*, lever/source), confidence
# PLACEHOLDERS: ERmax = 1, P* = 100 and ER* = the ad hoc ER at $100 (Manual inputs F40:F47), so beta equals the ad hoc
# value and results are unchanged. Replace with the values from the separate IPCC research report.
_PH = "PLACEHOLDER (ad hoc ER at $100); replace from IPCC research report. Lever: "
_LEVERS = {
    "DRI-EAF steel": ("biocarbon for charge/injection carbon; electrode efficiency", "n/a (no = 0)"),
    "Scrap-EAF steel": ("biocarbon for charge/injection carbon; electrode efficiency", "n/a (no = 0)"),
    "BF-BOF steel (reference)": ("flux (carbonate) reduction; deep options are process change/CCS", "n/a (no = 0)"),
    "Grey clinker (dry-process)": ("decarbonated raw materials; CCS excluded before 2035", "n/a (no = 0)"),
    "Ammonia (net merchant)": ("n/a (np = 0; SMR CO2 is fp)", "n/a (no = 0)"),
    "Urea": ("n/a (np < 0, captured CO2)", "n/a (no = 0)"),
    "Ammonium nitrate": ("n/a (np = 0)", "nitric-acid N2O: secondary/tertiary catalysts, NSCR"),
    "Primary aluminium": ("anode carbon consumption; inert anodes not commercial before 2035",
                          "PFC: anode-effect management, point feeders"),
}
_ER100 = (0.1, 0.02, 0.05, 0.08, 0.18, 0.0, 0.65, 0.08)     # Manual inputs F40:F47 (asserted by the builder)
PARAMS = {p: ((1.0, 100, e, _PH + _LEVERS[p][0]), (1.0, 100, e, _PH + _LEVERS[p][1]), "placeholder")
          for p, e in zip(PRODUCTS, _ER100)}

R_SEL, R_HDR, R0 = 50, 52, 53          # selector row, header row, first product row (Manual inputs)
HEAD = ("Product", "np ERmax", "np anchor P* ($/t)", "np ER at P*", "np beta (1/$)", "no ERmax",
        "no anchor P* ($/t)", "no ER at P*", "no beta (1/$)", "np semi-el. at P=0 (%/$)", "no semi-el. at P=0 (%/$)",
        "np ER at $20", "no ER at $20", "np lever and source", "no lever and source", "Confidence",
        "beta.np used", "ERmax.np used", "beta.no used", "ERmax.no used")   # D..W


def mi_section(wb):
    ws = wb.Worksheets("Manual inputs")
    for r in range(49, R0 + 9):
        if any(ws.Cells(r, c).Formula not in ("", None) for c in range(1, 32)):
            raise ValueError("Manual inputs row %d not empty" % r)
    if ws.Range("D40").Value != "DRI-EAF steel" or ws.Range("D47").Value != "Primary aluminium":
        raise ValueError("Manual inputs ad hoc semi-elasticity table moved")
    for k, p in enumerate(PRODUCTS):
        if ws.Range("D%d" % (30 + k)).Value != p:
            raise ValueError("Manual inputs EF row %d is not %s" % (30 + k, p))
        if abs(ws.Range("F%d" % (40 + k)).Value - _ER100[k]) > 1e-12:
            raise ValueError("ad hoc ER at $100 changed in Manual inputs F%d" % (40 + k))
    ws.Range("B38").Value = ("Legacy ad hoc semi-elasticities (AdHocCalculations; one beta for np + no; no documented "
                             "source). Used only when E%d = ADHOC, and by the fund placeholder (block o107-o114)" % R_SEL)
    copy_formats(ws.Rows(38), ws.Rows(49))
    copy_formats(ws.Rows(39), ws.Rows(R_HDR))
    for k in range(8):
        copy_formats(ws.Rows(40 + k), ws.Rows(R0 + k))
    wb.Application.CutCopyMode = False
    ws.Range("B49").Value = ("Process-emission semi-elasticities (Task D, v0.9): inferred from IPCC; used by CBAM block "
                             "rows o118-o125")
    ws.Range("D%d" % R_SEL).Value = "Semi-elasticity set used by the block (IPCC / ADHOC)"
    ws.Range("E%d" % R_SEL).Value = "IPCC"
    ws.Range("E%d" % R_SEL).Interior.Color = ws.Range("E21").Interior.Color
    v = ws.Range("E%d" % R_SEL).Validation
    v.Delete()
    v.Add(3, 1, 1, "IPCC,ADHOC")
    ws.Range("F%d" % R_SEL).Value = "ADHOC = rows 40-47 with ERmax = 1 (reproduces v0.8)"
    ws.Range("D%d" % (R_SEL + 1)).Value = (
        "ER(P) = ERmax x (1 - exp(-beta x P)); beta = -ln(1 - ER*/ERmax)/P* from the anchor (P*, ER*); semi-elasticity "
        "at P = 0 = ERmax x beta. np = non-fuel process CO2, no = non-CO2 process (Manual inputs J, K). ERmax = "
        "technical potential by ~2035 with commercial technology; CCS and process change are excluded (no Egyptian "
        "CO2 storage or H2-DRI before 2035).")
    for j, h in enumerate(HEAD):
        ws.Cells(R_HDR, 4 + j).Value = h
    ws.Range("D%d:W%d" % (R_HDR, R_HDR)).WrapText = True
    for k, p in enumerate(PRODUCTS):
        r = R0 + k
        (m1, p1, e1, s1), (m2, p2, e2, s2), conf = PARAMS[p]
        ws.Range("D%d" % r).Value = p
        ws.Range("E%d:G%d" % (r, r)).Value = (m1, p1, e1)
        ws.Range("I%d:K%d" % (r, r)).Value = (m2, p2, e2)
        ws.Range("H%d" % r).Formula = "=IF(OR(E{0}=0,G{0}=0),0,-LN(1-G{0}/E{0})/F{0})".format(r)
        ws.Range("L%d" % r).Formula = "=IF(OR(I{0}=0,K{0}=0),0,-LN(1-K{0}/I{0})/J{0})".format(r)
        ws.Range("M%d" % r).Formula = "=E{0}*H{0}".format(r)
        ws.Range("N%d" % r).Formula = "=I{0}*L{0}".format(r)
        ws.Range("O%d" % r).Formula = "=E{0}*(1-EXP(-H{0}*20))".format(r)
        ws.Range("P%d" % r).Formula = "=I{0}*(1-EXP(-L{0}*20))".format(r)
        ws.Range("Q%d:S%d" % (r, r)).Value = (s1, s2, conf)
        ws.Range("T%d" % r).Formula = '=IF($E${1}="ADHOC",$E${2},H{0})'.format(r, R_SEL, 40 + k)
        ws.Range("U%d" % r).Formula = '=IF($E${1}="ADHOC",1,E{0})'.format(r, R_SEL)
        ws.Range("V%d" % r).Formula = '=IF($E${1}="ADHOC",$E${2},L{0})'.format(r, R_SEL, 40 + k)
        ws.Range("W%d" % r).Formula = '=IF($E${1}="ADHOC",1,I{0})'.format(r, R_SEL)
        for c in "EFGIJK":
            ws.Range("%s%d" % (c, r)).Interior.Color = REVIEW if conf == "placeholder" else ws.Range("E21").Interior.Color
    last = R0 + 7
    ws.Range("E%d:G%d,I%d:K%d,O%d:P%d,U%d:U%d,W%d:W%d" % ((R0, last) * 5)).NumberFormat = "0%"
    ws.Range("F%d:F%d,J%d:J%d" % (R0, last, R0, last)).NumberFormat = "0"
    ws.Range("H%d:H%d,L%d:L%d,T%d:T%d,V%d:V%d" % ((R0, last) * 4)).NumberFormat = "0.00000"
    ws.Range("M%d:N%d" % (R0, last)).NumberFormat = "0.000%"
    for c in "QR":
        ws.Range("%s%d:%s%d" % (c, R0, c, last)).WrapText = False
    ws.Range("T%d:W%d" % (R_HDR, last)).Interior.Color = TAN


def update_block(ws, t0):
    if ws.Range("C%d" % (t0 + 117)).Value != "Emisssions Reductions, non-fuel process emissions":
        raise ValueError("process ER header moved (block %d)" % t0)
    hdr = (("E", "beta.np (1/$)"), ("F", "ERmax.np"), ("G", "beta.no (1/$)"), ("H", "ERmax.no"))
    for c, h in hdr:
        ws.Range("%s%d" % (c, t0 + 117)).Value = h
    for k, p in enumerate(PRODUCTS):
        r = t0 + 118 + k
        if ws.Range("D%d" % r).Value is None or ws.Range("G%d" % r).Formula or ws.Range("H%d" % r).Formula:
            raise ValueError("unexpected content in block row %d" % r)
        m = R0 + k
        for c, src in zip("EFGH", "TUVW"):
            ws.Range("%s%d" % (c, r)).Formula = "='Manual inputs'!%s$%d" % (src, m)
        ws.Range("E%d:E%d,G%d:G%d" % (r, r, r, r)).NumberFormat = "0.00000"
        ws.Range("F%d:F%d,H%d:H%d" % (r, r, r, r)).NumberFormat = "0%"
        price = "(R%dC6*R%dC+R%dC7*R%dC)" % (t0 + 65 + k, t0 + 63, t0 + 65 + k, t0 + 3)
        f = ("=-R%dC/1000*(R%dC10*MAX(R%dC10,0)*RC6*(1-EXP(-RC5*%s))+R%dC11*MAX(R%dC11,0)*RC8*(1-EXP(-RC7*%s)))"
             % (t0 + 31 + k, t0 + 28, t0 + 19 + k, price, t0 + 28, t0 + 19 + k, price))
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task D (v0.9): process-emission semi-elasticities (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:I%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff| / count", "", "",
                                              "Expected")
    MI = "'Manual inputs'!"
    M = "Mitigation_Industry!"
    a, z = R0, R0 + 7
    rng = lambda c: "%s$%s$%d:$%s$%d" % (MI, c, a, c, z)
    items = [
        ("Calibration identity: ERmax x (1 - exp(-beta x P*)) - ER* (np and no)",
         "=MAX(SUMPRODUCT(ABS(({E}>0)*({G}>0)*({E}*(1-EXP(-{H}*{F}))-{G}))),"
         "SUMPRODUCT(ABS(({I}>0)*({K}>0)*({I}*(1-EXP(-{L}*{J}))-{K}))))"),
        ("Parameter violations: ERmax outside [0,1], ER* >= ERmax, beta < 0, P* <= 0 with ER* > 0 (count)",
         "=SUMPRODUCT(({E}<0)+({E}>1)+({I}<0)+({I}>1)+({G}>0)*({G}>={E})+({K}>0)*({K}>={I})+({H}<0)+({L}<0)"
         "+({G}>0)*({F}<=0)+({K}>0)*({J}<=0))"),
        ("Selector %sE%d not IPCC or ADHOC (count)" % (MI, R_SEL),
         '=1-OR(%s$E$%d="IPCC",%s$E$%d="ADHOC")' % (MI, R_SEL, MI, R_SEL)),
        ("Effective parameters (T:W) differ from the selected set (max abs)",
         '=IF({S}="ADHOC",MAX(SUMPRODUCT(ABS({T}-{A})),SUMPRODUCT(ABS({V}-{A})),SUMPRODUCT(ABS({U}-1)),'
         'SUMPRODUCT(ABS({W}-1))),MAX(SUMPRODUCT(ABS({T}-{H})),SUMPRODUCT(ABS({V}-{L})),SUMPRODUCT(ABS({U}-{E})),'
         'SUMPRODUCT(ABS({W}-{I}))))'),
    ]
    sub = {c: rng(c) for c in "EFGHIJKLTUVW"}
    sub["S"] = "%s$E$%d" % (MI, R_SEL)
    sub["A"] = "%s$E$40:$E$47" % MI
    r = hdr + 1
    for item, f in items:
        ws.Range("A%d" % r).Value = item
        ws.Range("F%d" % r).Formula = f.format(**sub)
        ws.Range("I%d" % r).Value = 0
        r += 1
    for t0, scen in BLOCKS:
        s = 1 if t0 == BLOCKS[0][0] else 2
        b = lambda c, o: "%s$%s$%d:$%s$%d" % (M, c, t0 + o, c, t0 + o + 7)
        ws.Range("A%d" % r).Value = "Block E:H (o118-o125) differ from Manual inputs T:W (max abs)"
        ws.Range("D%d" % r).Value = s
        ws.Range("F%d" % r).Formula = (
            "=MAX(SUMPRODUCT(ABS(%s-%s)),SUMPRODUCT(ABS(%s-%s)),SUMPRODUCT(ABS(%s-%s)),SUMPRODUCT(ABS(%s-%s)))"
            % (b("E", 118), sub["T"], b("F", 118), sub["U"], b("G", 118), sub["V"], b("H", 118), sub["W"]))
        ws.Range("I%d" % r).Value = 0
        r += 1
        ws.Range("A%d" % r).Value = ("Process ER above 0 or larger than ERmax x covered pre-policy process emissions "
                                     "(count of product-years)")
        ws.Range("D%d" % r).Value = s
        ws.Range("F%d" % r).Formula = (
            "=SUMPRODUCT((%s$L$%d:$AE$%d>1E-12)+(-%s$L$%d:$AE$%d>%s$L$%d:$AE$%d*%s$J$%d:$J$%d"
            "*(%s$F$%d:$F$%d+%s$H$%d:$H$%d+ABS(%s$F$%d:$F$%d-%s$H$%d:$H$%d))/2+1E-9))"
            % (M, t0 + 118, t0 + 125, M, t0 + 118, t0 + 125, M, t0 + 96, t0 + 103, M, t0 + 85, t0 + 92,
               M, t0 + 118, t0 + 125, M, t0 + 118, t0 + 125, M, t0 + 118, t0 + 125, M, t0 + 118, t0 + 125))
        ws.Range("I%d" % r).Value = 0
        r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task D)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task D (stream 1, v0.9): process-emission semi-elasticities by product and category (np, no) in Manual "
        "inputs rows %d-%d (selector E%d: IPCC default, ADHOC = v0.8). Stream 2 (fund, Task G/J) should apply its own "
        "shadow price through the same parameters, not rows 40-47." % (R0, R0 + 7, R_SEL))


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.9"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.9: Task D section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.9"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task D: IPCC-inferred process semi-elasticities by product and category (np, no) with technical potential "
        "ERmax (Manual inputs rows %d-%d); block o118-o125 apply them per category. Selector Manual inputs E%d "
        "(IPCC default; ADHOC reproduces v0.8). Anchors are placeholders equal to the ad hoc values (results "
        "unchanged) pending the IPCC research report." % (R0, R0 + 7, R_SEL))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(DST, 0, False)
        xl.Calculation = -4135
        if wb.Worksheets("Settings").Range("B10").Value != "LEGACY":
            raise ValueError("source must be saved with Settings!B10 = LEGACY")
        mi_section(wb)
        ws = wb.Worksheets("Mitigation_Industry")
        for t0, scen in BLOCKS:
            update_block(ws, t0)
        chk = update_check(wb)
        update_scenarios(wb)
        update_settings(wb, chk)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        for b in list(xl.Workbooks):
            b.Close(False)
        xl.Quit()
        os.remove(DST)
        raise
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
    print("saved", DST)


if __name__ == "__main__":
    main()
