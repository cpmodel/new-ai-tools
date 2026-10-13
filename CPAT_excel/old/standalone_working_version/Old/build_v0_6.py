"""Build CPAT_Industry_Kernel_Egypt_v0.6.xlsx from v0.5 (Task B: 4-way emission-factor split with coverage switches).

Emission-factor categories (Manual inputs H30:K37, mirrored in CBAM block o19-o26 H:K):
  fc = fuel combustion CO2, fp = fuel used in the process / reduction CO2, np = non-fuel process CO2,
  no = non-CO2 process (N2O, PFC). fc + fp form the "fuel" group, priced with the fuel (energy CO2) inclusion flags
  (o54-o61); np + no form the "process" group, priced with the process inclusion flags (o65-o72).

Changes (both CBAM blocks; the 156-row block size is unchanged):
  * Scenarios bundle table: coverage switches c.fc / c.fp / c.np / c.no (columns M:P, 0/1, default 1 for every
    bundle; EgyptResultsInitial treats fc + fp as the "fuel combustion" share covered by upstream levies, 44% of
    CBAM-sector emissions, and np + no as "all process emissions"). Active values in row 27.
  * Block spare rows o27 (header) / o28 (switches, aligned with the EF columns H:K); scenario 1 uses 1.
  * Price rows: o54-o61 E = covered fuel EF = fc*c.fc + fp*c.fp; o65-o72 E = covered process EF
    = max(np,0)*c.np + max(no,0)*c.no. o85-o92 G (full process EF) = max(np,0) + max(no,0) (per-category floor;
    identical to legacy max(np+no,0) for current inputs: only urea np < 0, and its no = 0).
  * o85-o92 I / J: covered share of the fuel / process EF. Process ER (o118-o125) acts on the covered share only;
    revenues (o137, o138, o147-o154) use covered post-policy emissions. Physical emissions are unchanged.
  * Spare rows o13-o16: post-policy emissions by category, codes emisfc / emisfp / emisnp / emisno.cbam.tot
    (sum = emis.cbam.tot, o135).

With all switches = 1 (every current bundle) all values are identical to v0.5.

Run with Excel installed:  python build_v0_6.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLOCKS, BLUE, DATA_COLS, GREEN, REVIEW, TAN, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.5.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.5.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.6.xlsx")

CATS = ("fc", "fp", "np", "no")
CAT_COL = {"fc": 8, "fp": 9, "np": 10, "no": 11}           # H:K in block o19-o26 and o27-o28
CAT_LABEL = {
    "fc": "fuel combustion CO2",
    "fp": "fuel used in the process / reduction CO2",
    "np": "non-fuel process CO2 (floored at 0)",
    "no": "non-CO2 process (N2O, PFC; floored at 0)",
}
SW_HDR = ["%s covered (0/1)" % c for c in CATS]
SW_COLS = range(13, 17)                                     # Scenarios M:P
CODE = 'Settings!R3C2&".mit.%s."&RC7&"."&%s'                # % (variable, scenario-cell)


def empty(ws, r, c1=1, c2=31):
    return all(ws.Cells(r, c).Formula in ("", None) for c in range(c1, c2 + 1))


def update_block(ws, t0, scen):
    for off in (13, 14, 15, 16, 27, 28):
        if not empty(ws, t0 + off):
            raise ValueError("spare block row o%d (row %d) not empty" % (off, t0 + off))
    for off in range(84, 93):
        if not empty(ws, t0 + off, 9, 10):
            raise ValueError("block row o%d columns I:J not empty" % off)
    for k in range(8):
        r = t0 + 19 + k
        for c, cat in zip("HIJK", CATS):
            if "'Manual inputs'!%s$%d" % (c, 30 + k) not in ws.Range("%s%d" % (c, r)).Formula:
                raise ValueError("EF table o%d column %s does not read Manual inputs" % (19 + k, c))
        if ws.Range("E%d" % (t0 + 54 + k)).Formula != "=F%d" % (t0 + 85 + k) or \
                ws.Range("E%d" % (t0 + 65 + k)).Formula != "=G%d" % (t0 + 85 + k):
            raise ValueError("price-row EF links changed at o%d / o%d" % (54 + k, 65 + k))

    R = lambda o: "R%d" % (t0 + o)
    p, e, fund = R(2) + "C", R(3) + "C", R(4) + "C"
    sw = lambda cat: "%sC%d" % (R(28), CAT_COL[cat])

    # 1. coverage switches o27 (header) / o28 (values), aligned with the EF columns H:K
    copy_formats(ws.Rows(t0 + 7), ws.Rows(t0 + 27))
    ws.Application.CutCopyMode = False
    ws.Range("C%d" % (t0 + 27)).Value = "Emission-factor category coverage (1 = priced; Task B)"
    ws.Range("D%d" % (t0 + 28)).Value = ("Active bundle (Scenarios M:P); fc, fp priced with fuel flags (o54), "
                                         "np, no with process flags (o65); scenario 1 = 1")
    for cat in CATS:
        c = CAT_COL[cat]
        ws.Cells(t0 + 27, c).Value = "c." + cat
        ws.Cells(t0 + 28, c).FormulaR1C1 = (
            '=IF(%s>1,INDEX(Scenarios!R27C2:R27C26,MATCH("%s covered (0/1)",Scenarios!R26C2:R26C26,0)),1)'
            % (scen, cat))
        ws.Cells(t0 + 28, c).Interior.Color = TAN

    # 2. covered EFs in the price rows, full process EF floored per category, covered shares
    ws.Range("E%d" % (t0 + 53)).Value = "Covered fuel EF = fc*c.fc + fp*c.fp"
    ws.Range("E%d" % (t0 + 64)).Value = "Covered process EF = np*c.np + no*c.no (each >= 0)"
    ws.Range("I%d" % (t0 + 84)).Value = "Covered share, fuel (o54 E / F)"
    ws.Range("J%d" % (t0 + 84)).Value = "Covered share, process (o65 E / G)"
    for k in range(8):
        ws.Range("E%d" % (t0 + 54 + k)).FormulaR1C1 = "=R[-35]C8*%s+R[-35]C9*%s" % (sw("fc"), sw("fp"))
        ws.Range("E%d" % (t0 + 65 + k)).FormulaR1C1 = ("=MAX(R[-46]C10,0)*%s+MAX(R[-46]C11,0)*%s"
                                                        % (sw("np"), sw("no")))
        r = t0 + 85 + k
        ws.Range("G%d" % r).FormulaR1C1 = "=MAX(R[-66]C10,0)+MAX(R[-66]C11,0)"
        ws.Range("I%d" % r).FormulaR1C1 = "=IF(RC6=0,0,R[-31]C5/RC6)"
        ws.Range("J%d" % r).FormulaR1C1 = "=IF(RC7=0,0,R[-20]C5/RC7)"
    rng = ws.Range("I%d:J%d" % (t0 + 85, t0 + 92))
    rng.Interior.Color = TAN
    rng.NumberFormat = "0.000"

    # 3. process ER on the covered share only; revenues on covered post-policy emissions
    for off in range(118, 126):
        for c in DATA_COLS:
            ws.Cells(t0 + off, c).FormulaR1C1 = (
                "=R[-22]C*R[-33]C10*(EXP(-(RC5)*(R[-53]C6*%s+R[-53]C7*%s))-1)" % (p, e))
    for off in range(147, 155):
        for c in DATA_COLS:
            ws.Cells(t0 + off, c).FormulaR1C1 = (
                "=(R[-62]C+R[-40]C)*R[-62]C9*(R[-93]C6*%s+R[-93]C7*%s)"
                "+(R[-51]C*R[-62]C10+R[-29]C)*(R[-82]C6*%s+R[-82]C7*%s)" % (p, e, p, e))
    rng8 = lambda o, c="C": "%s%s:%s%s" % (R(o), c, R(o + 7), c)
    price = lambda fl: "(%s*%s+%s*%s)" % (rng8(fl, "C6"), p, rng8(fl, "C7"), e)
    for c in DATA_COLS:
        ws.Cells(t0 + 137, c).FormulaR1C1 = "=SUMPRODUCT((%s+%s)*%s*%s)" % (
            rng8(85), rng8(107), rng8(85, "C9"), price(54))
        ws.Cells(t0 + 138, c).FormulaR1C1 = "=SUMPRODUCT((%s*%s+%s)*%s)" % (
            rng8(96), rng8(85, "C10"), rng8(118), price(65))
    ws.Range("D%d" % (t0 + 137)).Value = ("Revenues, fuel CO2, covered fc + fp (memo: same CO2 is priced in kernel "
                                          "energy via nce)")
    ws.Range("D%d" % (t0 + 138)).Value = "Revenues, covered non-fuel process np + no (additional to energy carbon revenue)"

    # 4. post-policy emissions by EF category in spare rows o13-o16
    copy_formats(ws.Rows(t0 + 135), ws.Rows("%d:%d" % (t0 + 13, t0 + 16)))
    ws.Application.CutCopyMode = False
    ws.Range("C%d" % (t0 + 13)).Value = "Post-policy emissions by EF category (Task B)"
    prod = rng8(31)
    fuel_f = lambda cat: "=SUMPRODUCT(%s*%s*EXP(-%s*%s))/1000" % (
        prod, rng8(19, "C%d" % CAT_COL[cat]), rng8(107, "C5"), fund)

    def proc_f(cat):
        ef = rng8(19, "C%d" % CAT_COL[cat])
        return ("=SUMPRODUCT(%s*(%s>0)*%s*(1+%s*(EXP(-%s*%s)-1)))/1000"
                % (prod, ef, ef, sw(cat), rng8(118, "C5"), price(65)))
    for k, cat in enumerate(CATS):
        r = t0 + 13 + k
        ws.Range("D%d:E%d" % (r, r)).Value = ("%s: %s" % (cat, CAT_LABEL[cat]), "MtCO2e")
        ws.Range("G%d" % r).Value = "cbam.tot"
        ws.Range("H%d" % r).FormulaR1C1 = "=" + CODE % ("emis" + cat, scen)
        ws.Range("H%d" % r).Interior.Color = BLUE
        f = fuel_f(cat) if cat in ("fc", "fp") else proc_f(cat)
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = f


def update_scenarios(wb):
    ws = wb.Worksheets("Scenarios")
    if ws.Range("G16").Value != "Process emissions covered (0/1)" or ws.Range("B27").Formula != "=Settings!$B$10":
        raise ValueError("Scenarios bundle table layout changed")
    for r in list(range(16, 24)) + [26, 27]:
        if not all(ws.Cells(r, c).Formula in ("", None) for c in SW_COLS):
            raise ValueError("Scenarios M:P not empty at row %d" % r)
    copy_formats(ws.Range("G16"), ws.Range("M16:P16"))
    copy_formats(ws.Range("G17:G23"), ws.Range("M17:P23"))
    copy_formats(ws.Range("G26"), ws.Range("M26:P26"))
    copy_formats(ws.Range("G27"), ws.Range("M27:P27"))
    wb.Application.CutCopyMode = False
    for j, c in enumerate(SW_COLS):
        L = col(c)
        ws.Cells(16, c).Value = SW_HDR[j]
        ws.Cells(26, c).Value = SW_HDR[j]
        for r in range(17, 24):
            ws.Cells(r, c).Value = 1
        ws.Cells(27, c).Formula = "=INDEX(%s$17:%s$23,MATCH($B$27,$B$17:$B$23,0))" % (L, L)
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task B (stream 1): EF category coverage switches M:P (c.fc, c.fp, c.np, c.no; 1 in every bundle: the report "
        "treats fc + fp as the fuel-combustion share covered by energy CO2 pricing and np + no as all process emissions). "
        "fc, fp are priced with the energy CO2 flags, np, no with the process flags (E21). Set c.fp = 0 to exclude "
        "reductant/feedstock fuel, c.no = 0 to exclude non-CO2. Outputs emisfc / emisfp / emisnp / emisno.cbam.tot "
        "(block o13-o16); fuel switches affect only CBAM-block memo revenue, not the kernel energy wedge (nce).")


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task B (v0.6): EF category split and coverage switches (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff|", "", "", "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    M = "Mitigation_Industry!"
    r = hdr + 1
    for t0, scen in BLOCKS:
        s = 1 if t0 == BLOCKS[0][0] else 2
        ws.Range("A%d" % r).Value = "emis o135 - sum of EF categories o13-o16"
        ws.Range("D%d:E%d" % (r, r)).Value = (s, "diff")
        for c in DATA_COLS:
            L = col(c)
            ws.Cells(r, c).Formula = "=%s%s%d-SUM(%s%s%d:%s%d)" % (M, L, t0 + 135, M, L, t0 + 13, L, t0 + 16)
        ws.Range("F%d" % r).Formula = "=MAX(MAX(L%d:AE%d),-MIN(L%d:AE%d))" % (r, r, r, r)
        ws.Range("I%d" % r).Value = 0
        r += 1
        for item, f in (
                ("Covered EF above full EF (count of products, fuel + process)",
                 "=SUMPRODUCT((%s$E$%d:$E$%d>%s$F$%d:$F$%d+1E-12)+(%s$E$%d:$E$%d>%s$G$%d:$G$%d+1E-12))"
                 % (M, t0 + 54, t0 + 61, M, t0 + 85, t0 + 92, M, t0 + 65, t0 + 72, M, t0 + 85, t0 + 92)),
                ("Block switches o28 H:K not 0/1 (count)",
                 "=SUMPRODUCT((%s$H$%d:$K$%d<>0)*(%s$H$%d:$K$%d<>1))" % (M, t0 + 28, t0 + 28, M, t0 + 28, t0 + 28))):
            ws.Range("A%d" % r).Value = item
            ws.Range("D%d" % r).Value = s
            ws.Range("F%d" % r).Formula = f
            ws.Range("I%d" % r).Value = 0
            r += 1
    ws.Range("A%d" % r).Value = "Scenarios bundle switches M17:P23 not 0/1 (count)"
    ws.Range("F%d" % r).Formula = "=SUMPRODUCT((Scenarios!$M$17:$P$23<>0)*(Scenarios!$M$17:$P$23<>1))"
    ws.Range("I%d" % r).Value = 0
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task B)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_settings(wb, chk_row):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.6"
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.6: Task B section (row %d)." % chk_row
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.6"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task B: 4-way EF split (fc, fp, np, no) with coverage switches per bundle (Scenarios M:P, all 1; block o28). "
        "Price rows use covered EFs (fuel group fc+fp, process group np+no); process ER and revenues act on the covered "
        "share (o85-o92 I:J). Full process EF floored per category. Post-policy emissions by category in o13-o16 "
        "(emisfc/emisfp/emisnp/emisno.cbam.tot). All values identical to v0.5 with switches = 1.")
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
        update_scenarios(wb)
        ws = wb.Worksheets("Mitigation_Industry")
        for t0, scen in BLOCKS:
            update_block(ws, t0, scen)
        chk = update_check(wb)
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
