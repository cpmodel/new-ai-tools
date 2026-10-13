"""Build CPAT_Industry_Kernel_Egypt_v0.16.xlsx from v0.15 (final-results run 2; decisions agreed 2026, see CAVEATS.md).

D1  P* deflation: Manual inputs rows 53-60 F / J (P*) 100 -> 122. The IPCC AR6 ER* are at USD2019 100/t; the kernel
    carbon prices are USD2024, so USD2019 100 ~ USD2024 122 (US GDP deflator). beta = -ln(1 - ER*/ERmax) / 122.
D2  Fuel-CO2 reallocation in the national composition (Table2_Industry rows 59 / 60): where the CBAM block fuel CO2
    of a kernel sector exceeds the sector's CPAT energy CO2 (cement; nfm and mch have CPAT sector energy CO2 = 0),
    the block level is used and the excess is taken out of CPAT's non-kernel industry (nk1 = indx - ek1'), so the
    industry total is unchanged. Policy: block level x sector response rate where the sector has CPAT energy, else
    the block's own response (IPPU_Industry C1 / C2 reconciliation rows).
D3  EG3 full-coverage approximation (bundles 3A-3C, energy scope 'Industry only'): CPAT run EG3 prices only
    kappa ~ 0.54 of industry energy CO2. kappa = (effective / headline carbon price) x energy CO2 / industry energy CO2
    of the run (Table2_Industry row 54; 1 for 'All sectors'). The non-kernel industry response becomes
    nk1 x dInd / (Ind1 x kappa); J covers all industry energy CO2; P prices all of it; AR scales (dRev - receipts) by
    1 / kappa. K needs no other change (the 1 / kappa terms of dGHG, dIPPU and dInd cancel in K). Approximation.
D4  Abatement fund cap (bundle 3C): F = phi x R*, R* = block revenue after the fund response (o155) at the fixed
    point R* = o155(phi x R*). R* is stored per bundle (Fund_Industry rows 425-433) and solved here by a secant
    iteration per year; Check 'v0.16' asserts |phi x o155 - F| = 0 for every bundle. Previously F = phi x rev0
    (pre-fund revenue), which over-committed the fund and made 3C net revenue negative.
Snapshot (Table2_Industry section E) and the Task J v2 stored table refilled; rebuild column -> AdHoc rebuild v0.3.

Run with Excel installed (from this folder):  python build_v0_16.py
"""
import datetime
import json
import os
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))  # earlier builders archived in Old/

from build_v0_4 import DATA_COLS, GREEN, REVIEW, col, copy_formats
from build_v0_14 import BUNDLES, snapshot
from build_v0_15 import refill_v2, verify

SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.15.xlsx")
DST = os.environ.get("V016_DST") or os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.16.xlsx")
VER = "v0.16"
MIS = "'Manual inputs'!"
PSTAR = 122
RK_JSON = os.path.join(HERE, "..", "..", "..", "egypt", "supporting", "AdHocRebuild", "rebuild_v0_3_K2030.json")
REBUILD_K = (json.load(open(RK_JSON)) if os.path.exists(RK_JSON) else
             {"1A": -31.79, "2A": -27.82, "2B": -29.34, "3A": -18.15, "3B": -12.74, "3C": -22.62})
RK_LABEL = "Rebuild v0.3 K" if os.path.exists(RK_JSON) else "Rebuild v0.2 K"
L0, L1 = col(DATA_COLS[0]), col(DATA_COLS[-1])
YC = [col(c) for c in DATA_COLS]
# IPPU_Industry reconciliation rows: (block s1, sector s1, block s2, sector s2)
REC = ((51, 52, 75, 76), (56, 57, 80, 81), (61, 62, 85, 86), (66, 67, 90, 91))
FS0 = 425  # Fund_Industry stored R* block: title 425, header 426, bundles 427-433


def write_pstar(wb):
    mi = wb.Worksheets("Manual inputs")
    for r in range(53, 61):
        mi.Range("F%d" % r).Value = PSTAR
        mi.Range("J%d" % r).Value = PSTAR
        for c in ("Q", "R"):
            v = str(mi.Range("%s%d" % (c, r)).Value or "")
            mi.Range("%s%d" % (c, r)).Value = v.replace(
                "(Option A: ERmax 1, P* 100)",
                "(Option A: ERmax 1, P* %d = USD2019 100 deflated to USD2024, US GDP deflator ~1.22; v0.16)" % PSTAR)
    ck = wb.Worksheets("Check")
    n = 0
    for r in range(1880, ck.UsedRange.Row + ck.UsedRange.Rows.Count):
        a = ck.Range("A%d" % r).Value
        f = ck.Range("F%d" % r).Formula
        if isinstance(a, str) and ("ER(100)" in a or "ER(P*) used" in a) and "*100))" in f:
            ck.Range("F%d" % r).Formula = f.replace("*100))", "*%d))" % PSTAR)
            ck.Range("A%d" % r).Value = a.replace("ER(100)", "ER(%d)" % PSTAR) + " [P* %d, v0.16]" % PSTAR
            n += 1
    if n != 9:
        raise ValueError("expected 9 T1/T5 ER checks, found %d" % n)


def write_table2(wb):
    ws = wb.Worksheets("Table2_Industry")
    ip = "IPPU_Industry!"
    copy_formats(ws.Rows(53), ws.Rows(54))
    wb.Application.CutCopyMode = False
    ws.Range("B54:E54").Value = ("kap", "cpat", "kappa: share of industry energy CO2 priced by the run (Industry only: "
                                 "effective / headline price x energy CO2 / industry energy CO2; 1 for All sectors)",
                                 "share")
    ws.Range("F54").Value = "v0.16 EG3 full-coverage approximation (D3)"
    ws.Range("H54").Formula = '=Settings!$B$3&".mit.t2.kap"'
    for c in YC:
        f = {
            54: '=IF(OR($D$15="-",$D$16="All sectors"),1,IF(OR({c}33=0,{c}31=0),1,MIN(1,{c}34/{c}33*{c}29/{c}31)))',
            59: "=" + "+".join("MAX({p}{c}%d,{p}{c}%d)" % (b1, t1) for b1, t1, _b2, _t2 in REC),
            60: "=" + "+".join("IF({p}{c}%d>={p}{c}%d,{p}{c}%d,IF({p}{c}%d>0,{p}{c}%d*{p}{c}%d/{p}{c}%d,{p}{c}%d))"
                               % (t1, b1, t2, t1, b1, t2, t1, b2) for b1, t1, b2, t2 in REC),
            64: '=IF($D$15="-",0,IF(OR({c}31=0,{c}54=0),0,{c}63*{c}52/{c}31/{c}54))',
            84: '=IF({c}25=0,0,100*(IF($D$16="All sectors",{c}29,{c}72*{c}62)+$D$17*{c}69)/{c}25)',
            91: '=IF($D$15="-",0,IF($D$16="All sectors",{c}47+{c}71*{c}72*{c}82/1000,'
                '{c}71*{c}72*({c}60+{c}63+{c}64)/1000)+{c}73)',
            94: '=IF({c}51+(1/{c}54-1)*{c}52=0,0,{c}48/{c}54*{c}93/({c}51+(1/{c}54-1)*{c}52))',
            96: '=IF($D$15="-",0,{c}53-{c}47/{c}54+{c}92)',
        }
        for r, t in f.items():
            ws.Range("%s%d" % (c, r)).Formula = t.format(c=c, p=ip)
    ws.Range("D59").Value = ("Energy CO2, kernel sectors, baseline: sum over sectors of max(block fuel CO2, CPAT sector "
                             "energy CO2) (v0.16 reallocation; excess comes out of non-kernel industry)")
    ws.Range("D60").Value = ("Energy CO2, kernel sectors, policy: sector policy CO2 where sector >= block, else block "
                             "baseline x sector response rate (block's own response where the sector has 0)")
    ws.Range("F59").Value = "IPPU_Industry C1 rows 51-67 (v0.16 D2)"
    ws.Range("F60").Value = "IPPU_Industry C2 rows 75-91 (v0.16 D2)"
    ws.Range("D64").Value = ("dNon-kernel industry energy CO2 = nk1 x dInd CPAT / (Ind CPAT baseline x kappa) "
                             "(v0.16: Industry-only runs scaled to full industry coverage)")
    ws.Range("D84").Value = ("[J] Coverage, % of baseline GHG (All: energy CO2; Industry only: all industry energy CO2; "
                             "+ flag x block process)")
    ws.Range("D91").Value = ("Gross carbon revenue (All: CPAT receipts adjusted to the kernel industry; Industry only: "
                             "price x all industry energy CO2 incl. non-kernel at full coverage) + block process revenue")
    ws.Range("D94").Value = ("[Q] Air-pollution deaths avoided (CPAT deaths / kappa x dEnergy CO2 kernel / CPAT "
                             "dEnergy CO2 at full coverage, dE + (1/kappa - 1) dInd; kappa = 1 for All sectors)")
    ws.Range("D96").Value = "[AR] Net new revenue change (CPAT dRev - CPAT carbon-tax receipts / kappa + P)"
    ck = wb.Worksheets("Check")
    ck.Range("A1866").Value = ("Industry replacement = dEnergy kernel sectors + dnk - dInd CPAT (v0.16: non-kernel "
                               "industry priced at full coverage, kappa)")
    for c in YC:
        ck.Range("%s1866" % c).Formula = (
            '=IF(Table2_Industry!$D$15="-",0,Table2_Industry!{c}82-(Table2_Industry!{c}61+Table2_Industry!{c}64'
            '-Table2_Industry!{c}52))').format(c=c)


def write_fund(wb):
    fd = wb.Worksheets("Fund_Industry")
    copy_formats(fd.Rows(325), fd.Rows(FS0 + 1))
    for i in range(len(BUNDLES)):
        copy_formats(fd.Rows(328), fd.Rows(FS0 + 2 + i))
    wb.Application.CutCopyMode = False
    fd.Range("B%d" % FS0).Value = ("Stored fixed point (v0.16, build_v0_16.py): R* = block revenue after the fund "
                                   "response (o155) solving R* = o155(phi x R*), per bundle, scenario 2")
    fd.Range("B%d" % FS0).Font.Bold = True
    fd.Range("B%d:E%d" % (FS0 + 1, FS0 + 1)).Value = ("bundle", "Item", "Description", "Unit")
    for i, bd in enumerate(BUNDLES):
        r = FS0 + 2 + i
        fd.Range("B%d:E%d" % (r, r)).Value = (bd, "cbam", "Post-fund block revenue R* (stored; refreshed by the "
                                              "builder; Check v0.16 flags a stale value)", "USD million")
        fd.Range("H%d" % r).Value = ""
        fd.Range("%s%d:%s%d" % (L0, r, L1, r)).Interior.Color = REVIEW
    fd.Range("D329").Value = "Abatement fund F = phi x R* (post-fund block revenue, fixed point; rows %d-%d)" % (
        FS0 + 2, FS0 + 1 + len(BUNDLES))
    for c in YC:
        fd.Range("%s329" % c).Formula = "=%s$218*INDEX(%s$%d:%s$%d,MATCH(Settings!$B$10,$B$%d:$B$%d,0))" % (
            c, c, FS0 + 2, c, FS0 + 1 + len(BUNDLES), FS0 + 2, FS0 + 1 + len(BUNDLES))


def solve_fund(wb, tol=1e-9, itmax=40):
    xl, st, fd, mi = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Fund_Industry"), \
        wb.Worksheets("Mitigation_Industry")
    rng = lambda r: "%s%d:%s%d" % (L0, r, L1, r)
    log = {}
    for i, bd in enumerate(BUNDLES):
        r = FS0 + 2 + i
        st.Range("B10").Value = bd
        xl.CalculateFull()
        if any(abs(x or 0) > 1e-12 for x in fd.Range(rng(122)).Value[0]):
            raise ValueError("scenario-1 fund non-zero for %s (cap not implemented for scenario 1)" % bd)
        phi = fd.Range(rng(218)).Value[0]
        x0 = list(fd.Range(rng(328)).Value[0])  # rev0 = o155 at F = 0

        def g(x):
            fd.Range(rng(r)).Value = [x]
            xl.CalculateFull()
            o = mi.Range(rng(799)).Value[0]
            return [a - b for a, b in zip(x, o)], o
        g0, o0 = g(x0)
        if max(phi) == 0:
            log[bd] = (0, 0.0)
            continue
        x1 = list(o0)
        g1, o1 = g(x1)
        it = 1
        while max(abs(v) for v in g1) > tol and it < itmax:
            x2 = [b - gb * (b - a) / (gb - ga) if abs(gb - ga) > 1e-15 else ob
                  for a, b, ga, gb, ob in zip(x0, x1, g0, g1, o1)]
            x0, g0 = x1, g1
            x1 = x2
            g1, o1 = g(x1)
            it += 1
        if max(abs(v) for v in g1) > tol:
            raise ValueError("fund fixed point not converged for %s: %g" % (bd, max(abs(v) for v in g1)))
        log[bd] = (it, max(abs(v) for v in g1))
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return log


def t2_rows(wb):
    ws = wb.Worksheets("Table2_Industry")
    k = {"snap": []}
    for r in range(1, ws.UsedRange.Rows.Count + 1):
        a = ws.Range("A%d" % r).Value
        if a == "live":
            k["live"] = r
        elif a == "stored":
            k["snap"].append(r)
    hdr = k["live"] - 1
    k["kc"] = next(c for c in range(1, 40) if str(ws.Cells(hdr, c).Value).startswith("Rebuild v0."))
    k["n"] = k["kc"] - 6
    return k


def refill_table2(wb, k):
    xl, st, ws = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Table2_Industry")
    hdr = k["live"] - 1
    ws.Cells(hdr, k["kc"]).Value = RK_LABEL
    for r in range(k["snap"][-1] + 1, k["snap"][-1] + 4):
        v = ws.Range("B%d" % r).Value
        if isinstance(v, str) and "Rebuild v0.2" in v:
            ws.Range("B%d" % r).Value = v.replace("build_v0_15.py", "build_v0_16.py").replace(
                "Rebuild v0.2", RK_LABEL[:-2])
    last = col(5 + k["n"])
    for r in k["snap"]:
        bd = ws.Range("B%d" % r).Value
        ws.Cells(r, k["kc"]).Value = REBUILD_K[bd]
        st.Range("B10").Value = bd
        xl.CalculateFull()
        ws.Range("C%d:E%d" % (r, r)).Value = ws.Range("C%d:E%d" % (k["live"], k["live"])).Value
        ws.Range("F%d:%s%d" % (r, last, r)).Value = ws.Range("F%d:%s%d" % (k["live"], last, k["live"])).Value
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    ck = wb.Worksheets("Check")
    for r in range(1, ck.UsedRange.Rows.Count + 1):
        v = ck.Range("A%d" % r).Value
        if isinstance(v, str) and "TASK-2b rebuild v0.2" in v and RK_LABEL.endswith("v0.3 K"):
            ck.Range("A%d" % r).Value = v.replace("rebuild v0.2", "rebuild v0.3")


def add_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "v0.16: P* deflation, fuel-CO2 reallocation, EG3 kappa scaling, fund fixed point"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 3
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:I%d" % (hdr, hdr)).Value = ("Item", "", "", "", "Kind", "|diff| / count", "", "", "Expected")
    t2, fd, ip = "Table2_Industry!", "Fund_Industry!", "IPPU_Industry!"
    yr = "$%s$%%d:$%s$%%d" % (L0, L1)
    items = [
        ("P* = %d on every row with ERmax > 0 (np E/F, no I/J; max abs)" % PSTAR,
         "=SUMPRODUCT(('Manual inputs'!$E$53:$E$60>0)*ABS('Manual inputs'!$F$53:$F$60-%d))+SUMPRODUCT(("
         "'Manual inputs'!$I$53:$I$60>0)*ABS('Manual inputs'!$J$53:$J$60-%d))" % (PSTAR, PSTAR)),
        ("Fund fixed point: |phi x o155 - F| scenario 2, active bundle (USD million, max abs over years)",
         "=SUMPRODUCT(ABS(%s%s))" % (fd, yr % (407, 407))),
        ("Reallocated kernel energy CO2 ek1' below IPPU_Industry row 17 (count of years)",
         "=SUMPRODUCT(--(%s%s<%s%s-1E-9))" % (t2, yr % (59, 59), ip, yr % (17, 17))),
        ("Non-kernel industry energy CO2 nk1 < 0 after reallocation (count of years)",
         "=SUMPRODUCT(--(%s%s<-1E-9))" % (t2, yr % (63, 63))),
        ("Industry total preserved: ek1' + nk1 - CPAT industry energy CO2 (max abs)",
         "=SUMPRODUCT(ABS(%s%s+%s%s-%s%s))" % (t2, yr % (59, 59), t2, yr % (63, 63), t2, yr % (62, 62))),
        ("kappa outside (0, 1] (count of years)",
         "=SUMPRODUCT((%s%s<=0)+(%s%s>1))" % (t2, yr % (54, 54), t2, yr % (54, 54))),
    ]
    r = hdr + 1
    for label, f in items:
        ws.Range("A%d" % r).Value = label
        ws.Range("E%d" % r).Value = "diff"
        ws.Range("F%d" % r).Formula = f
        ws.Range("I%d" % r).Value = 0
        r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (v0.16)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=IF(MAX(F%d:F%d)<1E-6,0,MAX(F%d:F%d))" % (hdr + 1, r - 1, hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def fix_v2_note(wb):
    ck = wb.Worksheets("Check")
    tag = " [refilled by build_v0_15.py: phi = 0 per bundle, T1/T5 inputs]"
    for r in range(1, ck.UsedRange.Rows.Count + 1):
        v = ck.Cells(r, 1).Value
        if isinstance(v, str) and v.endswith(tag + tag):
            ck.Cells(r, 1).Value = v[:-2 * len(tag)] + " [refilled by build_v0_16.py: phi = 0 per bundle, v0.16 inputs]"


def update_notes(wb, chk, stats, fund_log):
    sc = wb.Worksheets("Scenarios")
    last = sc.Cells(sc.Rows.Count, 2).End(-4162).Row
    sc.Range("B%d" % (last + 1)).Value = (
        "v0.16: P* 122 (USD2024); fuel-CO2 reallocation in the national composition (Table2 rows 59/60); EG3 "
        "full-coverage approximation via kappa (Table2 row 54); fund F = phi x post-fund block revenue (fixed point, "
        "Fund_Industry rows %d-%d). Table 2 snapshot refilled." % (FS0 + 2, FS0 + 1 + len(BUNDLES)))
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: v0.16 section (row %d)." % (VER, chk)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Final-results run 2. D1 Manual inputs rows 53-60 P* 100 -> 122 (USD2019 100 in USD2024). D2 Table2 rows 59/60: "
        "kernel energy CO2 = max(block fuel CO2, CPAT sector) per sector, excess out of non-kernel industry (industry "
        "total unchanged). D3 Table2 row 54 kappa; Industry-only runs (EG3) scaled to full industry coverage (rows 64, "
        "84, 91, 96; approximation). D4 Fund_Industry row 329 F = phi x R*, R* stored rows %d-%d (secant fixed point, "
        "iterations %s). Check row %d. 2030 K: %s." % (FS0 + 2, FS0 + 1 + len(BUNDLES),
                                                     {b: v[0] for b, v in fund_log.items() if v[0]}, chk, stats))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(SRC, 0, True)
        xl.Calculation = -4135
        print("snapshot v0.15 ...")
        snap = snapshot(wb)
        k = t2_rows(wb)
        t2 = wb.Worksheets("Table2_Industry")
        old = {t2.Range("B%d" % r).Value: [t2.Range("%s%d" % (c, r)).Value for c in "FGHIJKL"] for r in k["snap"]}
        hdrs = [t2.Range("%s%d" % (c, k["live"] - 1)).Value for c in "FGHIJKL"]
        wb.SaveAs(DST)
        write_pstar(wb)
        write_table2(wb)
        write_fund(wb)
        xl.CalculateFull()
        print("v2 rows refilled:", refill_v2(wb))
        fix_v2_note(wb)
        fund_log = solve_fund(wb)
        print("fund fixed point:", fund_log)
        refill_table2(wb, k)
        chk = add_check(wb)
        xl.CalculateFull()
        print("verify ...")
        verify(wb, snap, k, chk)
        new = {t2.Range("B%d" % r).Value: [t2.Range("%s%d" % (c, r)).Value for c in "FGHIJKL"] for r in k["snap"]}
        print("columns", hdrs, "keys", list(new))
        for bd in new:
            print("%-6s v0.15 %s\n       v0.16 %s" % (bd, ["%.2f" % x if isinstance(x, float) else x for x in old.get(bd, [])],
                                                   ["%.2f" % x if isinstance(x, float) else x for x in new[bd]]))
        txt = ", ".join("%s %.1f" % (bd, new[bd][1]) for bd in new if isinstance(new[bd][1], float))
        update_notes(wb, chk, txt, fund_log)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        for b in list(xl.Workbooks):
            b.Close(False)
        xl.Quit()
        if os.path.exists(DST):
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
