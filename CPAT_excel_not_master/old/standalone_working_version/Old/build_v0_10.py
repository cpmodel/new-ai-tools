"""Build CPAT_Industry_Kernel_Egypt_v0.10.xlsx from v0.9 (Task H: output response to the net product cost increase).

Output (production) of each CBAM product responds to the carbon cost per tonne NET of the output-based rebate:

  dp_net = (fuel carbon cost o54 + process carbon cost o65
            - min(obrrb, cptraj) x c.fuel x EF.fuel - min(obrrb, pptraj) x c.proc x EF.proc) / pre-policy price o43
  Q      = Qbase x (1 + dp_net) ^ eps                     Qbase = Q2024 x (1 + g) ^ (year - 2024)

Intensity (ER rows) still sees the full price (+ fund shadow price o4); output sees price - obrrb (Scenarios contract).
With theta = 1 (bundle 3B, free allocation at a pre-policy intensity benchmark) dp_net = 0 and output is unchanged.

Manual inputs rows 62-73 (new section): switch E63 (ON / OFF; OFF reproduces v0.9) and the output elasticity eps per
product (E66:E73, PLACEHOLDER -0.5, review fill); effective eps in K66:K73 (0 when OFF or when the active bundle is
LEGACY, so the regression bundle still reproduces CPAT).

CBAM block (both scenarios):
  o30-o38 J      eps used (links to Manual inputs H66:H73)
  o31-o38 O:AE   production = baseline path x output factor (closed form; equals the v0.9 compounding path when eps = 0)
  o74-o82        redefined (were computed but unused): net cost increase as % of the pre-policy product price
  o84-o103       'pre-policy' emissions = pre-policy intensity x policy-scenario output, so Task K metrics (cbintch,
                 cbint, cbintchx) remain pure intensity measures ('same output')
  o62 cbqch      output change vs baseline output, weighted by baseline-output pre-policy emissions (share)
  o73 emrq       emissions change from the output response, at pre-policy intensity (MtCO2e, negative = reduction)
  o83 emrt       total change vs pre-policy emissions at baseline output = emrq + emr (o136)

Decomposition: emis (o135) = pre-policy emissions at baseline output + emrq + emr (checked at 2030 and 2041).
Fuel CO2 in the block is a memo (the kernel prices it via nce), so only the process part of emrq adds to national totals.

Run with Excel installed:  python build_v0_10.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, DATA_COLS, REVIEW, TAN, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.9.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.9.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.10.xlsx")

PRODUCTS = ("DRI-EAF steel", "Scrap-EAF steel", "BF-BOF steel (reference)", "Grey clinker (dry-process)",
            "Ammonia (net merchant)", "Urea", "Ammonium nitrate", "Primary aluminium")
EPS = -0.5          # PLACEHOLDER uniform output elasticity w.r.t. the net cost increase (% of price)
SRC_TXT = ("PLACEHOLDER uniform -0.5 (user choice, v0.10). Combines demand elasticity and limited pass-through for a "
           "trade-exposed commodity; replace with product-specific values (trade exposure, pass-through, leakage "
           "literature)")

R_TITLE, R_SW, R_NOTE, R_HDR, R0 = 62, 63, 64, 65, 66      # Manual inputs
HEAD = ("Product", "Output elasticity eps (input)", "Source / rationale", "Confidence", "", "", "", "eps used")  # D..K
FIRST_PROJ = 15     # column O = 2025, first projected year


def mi_section(wb):
    ws = wb.Worksheets("Manual inputs")
    for r in range(R_TITLE, R0 + 9):
        if any(ws.Cells(r, c).Formula not in ("", None) for c in range(1, 32)):
            raise ValueError("Manual inputs row %d not empty" % r)
    for k, p in enumerate(PRODUCTS):
        if ws.Range("D%d" % (53 + k)).Value != p:
            raise ValueError("Manual inputs row %d is not %s" % (53 + k, p))
    copy_formats(ws.Rows(38), ws.Rows(R_TITLE))
    copy_formats(ws.Rows(39), ws.Rows(R_HDR))
    for k in range(8):
        copy_formats(ws.Rows(40 + k), ws.Rows(R0 + k))
    wb.Application.CutCopyMode = False
    ws.Range("B%d" % R_TITLE).Value = ("Output response to the net product cost increase (Task H, v0.10): used by "
                                       "CBAM block production rows o31-o38")
    ws.Range("D%d" % R_SW).Value = "Output response switch (ON / OFF)"
    ws.Range("E%d" % R_SW).Value = "ON"
    ws.Range("E%d" % R_SW).Interior.Color = ws.Range("E21").Interior.Color
    v = ws.Range("E%d" % R_SW).Validation
    v.Delete()
    v.Add(3, 1, 1, "ON,OFF")
    ws.Range("F%d" % R_SW).Value = ("OFF = eps 0, production on the fixed growth path (reproduces v0.9). Always off "
                                    "for the LEGACY bundle (regression vs CPAT)")
    ws.Range("D%d" % R_NOTE).Value = (
        "Q = Qbase x (1 + dp_net)^eps; dp_net = (carbon cost per t - OBR rebate per t) / pre-policy product price "
        "(block o75-o82). The rebate obrrb (theta x cptraj) is paid on the pre-policy covered EF, so theta = 1 (3B) "
        "leaves output unchanged; intensity still responds to the full price.")
    for j, h in enumerate(HEAD):
        if h:
            ws.Cells(R_HDR, 4 + j).Value = h
    ws.Range("D%d:K%d" % (R_HDR, R_HDR)).WrapText = True
    for k, p in enumerate(PRODUCTS):
        r = R0 + k
        ws.Range("D%d" % r).Value = p
        ws.Range("E%d" % r).Value = EPS
        ws.Range("F%d" % r).Value = SRC_TXT
        ws.Range("G%d" % r).Value = "placeholder"
        ws.Range("K%d" % r).Formula = '=IF(AND($E${0}="ON",Settings!$B$10<>"LEGACY"),E{1},0)'.format(R_SW, r)
        ws.Range("E%d" % r).Interior.Color = REVIEW
        ws.Range("F%d:G%d" % (r, r)).WrapText = False
        ws.Range("F%d:J%d" % (r, r)).Interior.ColorIndex = -4142
    last = R0 + 7
    ws.Range("E%d:E%d,K%d:K%d" % (R0, last, R0, last)).NumberFormat = "0.00"
    ws.Range("K%d:K%d" % (R_HDR, last)).Interior.Color = TAN


def update_block(ws, t0):
    o = lambda n: t0 + n
    checks = ((o(29), 14, 2024), (o(74), 3, "Price Increase total"), (o(84), 3, "Pre-policy Fuel Emissions"),
              (o(95), 3, "Pre-policy Non-Fuel and Non-CO2 Process Emissions"))
    for r, c, v in checks:
        if ws.Cells(r, c).Value != v:
            raise ValueError("block %d row %d col %d is %r, expected %r" % (t0, r, c, ws.Cells(r, c).Value, v))
    for n in (62, 73, 83):
        if any(ws.Cells(o(n), c).Formula not in ("", None) for c in range(1, 36)):
            raise ValueError("spare row o%d not empty (block %d)" % (n, t0))
    # eps column and production
    ws.Range("J%d" % o(30)).Value = "Output elasticity eps used (Task H)"
    for k in range(8):
        r = o(31 + k)
        if ws.Range("J%d" % r).Formula not in ("", None):
            raise ValueError("block row %d col J not empty" % r)
        ws.Range("J%d" % r).Formula = "='Manual inputs'!K$%d" % (R0 + k)
        ws.Range("J%d" % r).NumberFormat = "0.00"
        f = "=RC14*(1+RC7)^(R%dC-R%dC14)*(1+R%dC)^RC10" % (o(29), o(29), o(75 + k))
        for c in DATA_COLS:
            if c >= FIRST_PROJ:
                ws.Cells(r, c).FormulaR1C1 = f
    # net cost increase (% of price)
    ws.Range("C%d" % o(74)).Value = ("Net carbon cost increase, % of pre-policy product price (fuel + process cost "
                                     "- OBR rebate on covered EF; drives output, Task H)")
    for k in range(8):
        r = o(75 + k)
        f = ("=(R{p}C+R{fu}C-MIN(R{rb}C,R{cp}C)*R{fu}C6*R{fu}C5-MIN(R{rb}C,R{pp}C)*R{p}C6*R{p}C5)/R{pr}C"
             .format(p=o(65 + k), fu=o(54 + k), rb=o(6), cp=o(2), pp=o(63), pr=o(43 + k)))
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f
    ws.Range("C%d" % o(84)).Value = "Pre-policy Fuel Emissions (pre-policy intensity x policy-scenario output, Task H)"
    ws.Range("C%d" % o(95)).Value = ("Pre-policy Non-Fuel and Non-CO2 Process Emissions (pre-policy intensity x "
                                     "policy-scenario output, Task H)")
    # metrics
    pre = "(R{a}C:R{b}C+R{c}C:R{d}C)".format(a=o(85), b=o(92), c=o(96), d=o(103))
    inv = "(1+R{a}C:R{b}C)^(-R{c}C10:R{d}C10)".format(a=o(75), b=o(82), c=o(31), d=o(38))
    rows = (
        (62, 41, "Output change vs baseline output, CBAM products (weighted by baseline-output pre-policy emissions; "
                 "Task H)", "share", "cbqch",
         "=IF(SUMPRODUCT({p}*{i})=0,0,R{e}C/SUMPRODUCT({p}*{i}))".format(p=pre, i=inv, e=o(73))),
        (73, 126, "Emissions change from the output response, at pre-policy intensity (Task H)", "MtCO2e", "emrq",
         "=SUMPRODUCT({p}*(1-{i}))".format(p=pre, i=inv)),
        (83, 126, "Total emissions change vs pre-policy emissions at baseline output (output emrq + intensity emr)",
         "MtCO2e", "emrt", "=R{a}C+R{b}C".format(a=o(73), b=o(136))),
    )
    scen_ref = ws.Range("H%d" % o(126)).Formula.rsplit("&", 1)[1]
    if not scen_ref.startswith("$J$"):
        raise ValueError("unexpected emrp code formula in block %d" % t0)
    for n, fmt_from, label, unit, code, f in rows:
        r = o(n)
        copy_formats(ws.Rows(o(fmt_from)), ws.Rows(r))
        ws.Range("D%d" % r).Value = label
        ws.Range("E%d" % r).Value = unit
        ws.Range("G%d" % r).Value = "cbam.tot"
        ws.Range("H%d" % r).Formula = '=Settings!$B$3&".mit.%s."&$G%d&"."&%s' % (code, r, scen_ref)
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f
    ws.Application.CutCopyMode = False


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task H (v0.10): output response to the net cost increase (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:I%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff| / count", "", "",
                                              "Expected")
    MI = "'Manual inputs'!"
    M = "Mitigation_Industry!"
    rng = lambda c: "%s$%s$%d:$%s$%d" % (MI, c, R0, c, R0 + 7)
    sw = "%s$E$%d" % (MI, R_SW)
    items = [
        ("Switch %sE%d not ON or OFF (count)" % (MI, R_SW), '=1-OR(%s="ON",%s="OFF")' % (sw, sw)),
        ("Effective eps (K) differs from IF(switch = ON and bundle <> LEGACY, E, 0) (max abs)",
         '=SUMPRODUCT(ABS(%s-(%s="ON")*(Settings!$B$10<>"LEGACY")*%s))' % (rng("K"), sw, rng("E"))),
        ("Output elasticity eps > 0 (count)", "=SUMPRODUCT(--(%s>0))" % rng("E")),
    ]
    r = hdr + 1
    for item, f in items:
        ws.Range("A%d" % r).Value = item
        ws.Range("F%d" % r).Formula = f
        ws.Range("I%d" % r).Value = 0
        r += 1
    for t0, scen in BLOCKS:
        s = 1 if t0 == BLOCKS[0][0] else 2
        b = lambda c, n, m=7: "%s$%s$%d:$%s$%d" % (M, c, t0 + n, c, t0 + n + m)
        add = []
        add.append(("Block J (o31-o38) differs from Manual inputs K (max abs)",
                    "=SUMPRODUCT(ABS(%s-%s))" % (b("J", 31), rng("K"))))
        add.append(("1 + net cost increase (o75-o82, L:AE) <= 0 (count)",
                    "=SUMPRODUCT(--(1+%s$L$%d:$AE$%d<=0))" % (M, t0 + 75, t0 + 82)))
        if s == 1:
            add.append(("Net cost increase non-zero in scenario 1 (sum abs; baseline output unchanged)",
                        "=SUMPRODUCT(ABS(%s$L$%d:$AE$%d))" % (M, t0 + 75, t0 + 82)))
        for yc in ("T", "AE"):
            base = ("SUMPRODUCT({N}*(1+{G})^({M}{y}${h}-{M}$N${h})*({F85}+{G85}))/1000"
                    .format(N=b("N", 31), G=b("G", 31), M=M, y=yc, h=t0 + 29, F85=b("F", 85), G85=b("G", 85)))
            add.append(("Decomposition emis = pre-policy at baseline output + emrq + emr, %s (abs)"
                        % ("2030" if yc == "T" else "2041"),
                        "=ABS({M}{y}{e}-{M}{y}{q}-{M}{y}{r}-{base})".format(
                            M=M, y=yc, e=t0 + 135, q=t0 + 73, r=t0 + 136, base=base)))
        for item, f in add:
            ws.Range("A%d" % r).Value = item
            ws.Range("D%d" % r).Value = s
            ws.Range("F%d" % r).Formula = f
            ws.Range("I%d" % r).Value = 0
            r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task H)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task H (stream 1, v0.10): CBAM production o31-o38 = baseline path x (1 + net cost increase o75-o82)^eps; the "
        "net cost deducts obrrb (o6, capped at the price paid) on the covered pre-policy EF. eps in Manual inputs "
        "rows %d-%d (switch E%d). Codes cbqch (o62), emrq (o73), emrt (o83). Stream 2 (Task I) should read output "
        "effects through these rows, not re-derive them." % (R0, R0 + 7, R_SW))


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.10"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.10: Task H section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.10"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task H: CBAM output responds to the carbon cost net of the output-based rebate, Q = Qbase x (1 + dp_net)^eps "
        "(Manual inputs rows %d-%d, switch E%d; eps placeholder -0.5). New codes cbqch, emrq, emrt; block o75-o82 "
        "redefined as the net cost increase. OFF (and the LEGACY bundle) reproduces v0.9." % (R0, R0 + 7, R_SW))
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
