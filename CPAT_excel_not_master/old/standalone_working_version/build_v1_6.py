"""Build CPAT_Industry_Kernel_Egypt_v1.6.xlsx from v1.5: product-specific output elasticities.

DRAFTED ON LINUX, NOT RUN. Needs Excel + pywin32 on Windows and kernel v1.5 in this folder. Run from this folder in a normal
(non-sandboxed) shell:   python build_v1_6.py

What changes (egypt/supporting/OutputElasticity_Note_v0.1.md): 'Manual inputs' E66:E73 (output elasticity eps) go from the
uniform -0.5 placeholder to cement -0.10, steel -0.40 (three routes), ammonia / urea / ammonium nitrate -0.40, aluminium -0.50.
Every scenario-2 result moves, so after setting the values this builder, exactly as build_v0_16.py did:
  1. refills the Task J stored 'v2' pre-fund table (build_v0_15.refill_v2);
  2. re-solves the 3C abatement-fund fixed point per bundle and refills Fund_Industry rows 427-433 (build_v0_16.solve_fund);
  3. refills the stored Table2_Industry snapshot rows (build_v0_16.refill_table2; the rebuild column is left as stored);
  4. refreshes the stored 2030 snapshot of CarveOut_Table2 (rows 80-85, columns C:L, and M = live intensity-only row O) for
     all six scenarios from the live sheet;
  5. rewrites the printed-figures block of Table2_Final (section B) from the Python mirror (make_carveout_v0_5.py) and recounts
     the mismatches; Settings title / version log; Check section.
Because the mirror is an approximation (3C fund shadow price not re-solved), the live results are exported to
carveout_v1_6_results_live.json. If the mismatch count is not 0 the builder still saves, prints
'DOCUMENTS MUST BE REGENERATED' and the documents are rebuilt from the live JSON (see TODO.md, Final steps).
Regression: nothing is expected to be unchanged in scenario 2; the LEGACY scenario (cpat-based) and baseline rows must be: the
builder verifies that every LEGACY-state value of every sheet except the intended areas equals v1.5 (max abs diff 0).
After a clean run it copies the kernel to egypt/final/ and moves the previous final kernel to egypt/archive/.
"""
import datetime
import json
import os
import shutil
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))
sys.path.insert(0, HERE)
import build_v0_15 as b15  # noqa: E402
import build_v0_16 as b16  # noqa: E402
from build_v0_4 import REVIEW, TAN, copy_formats  # noqa: E402

REPO = os.path.abspath(os.path.join(HERE, "..", ".."))
ADHOC = os.path.join(REPO, "egypt", "supporting", "AdHocRebuild")
FINAL = os.path.join(REPO, "egypt", "final")
ARCHIVE = os.path.join(REPO, "egypt", "archive")
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.5.xlsx")
SRC_OLD = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v1.5.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.6.xlsx")
VER = "v1.6"
BUNDLES = ["1A", "2A", "2B", "3A", "3B", "3C"]
CO, T2F, MI = "CarveOut_Table2", "Table2_Final", "Manual inputs"

EPS_NEW = [-0.40, -0.40, -0.40, -0.10, -0.40, -0.40, -0.40, -0.50]   # rows 66-73
PRODUCTS = ["DRI-EAF steel", "Scrap-EAF steel", "BF-BOF steel (reference)", "Grey clinker (dry-process)",
            "Ammonia (net merchant)", "Urea", "Ammonium nitrate", "Primary aluminium"]
BASIS = [
    "Steel: demand -0.2 to -0.3, pass-through about 0.5 (EU studies 0.06-1.2), import competition (GTAP Armington about 3). "
    "Replaces the uniform -0.5 placeholder; carbon cost is only 1-3% of the price, so the result is insensitive.",
] * 3 + [
    "Cement: market demand -0.02 to -0.16, pass-through 0.2-0.4 (EC / CE Delft-Oeko), small trade exposure (1.6% of output to "
    "the EU). The only product for which eps matters (carbon cost 15% of price).",
] + [
    "Fertiliser: inelastic demand; exporters are price takers. Carbon cost under 3% of the price, so the result is insensitive.",
] * 3 + [
    "Aluminium: world-priced metal, 56% exported; no better evidence found, kept at -0.5 (range -0.3 to -1.5). Cost under 1% "
    "of price: insensitive.",
]
METRICS = [("J", 0), ("P", 1), ("K", 1), ("L", 1), ("M", 0), ("N", 1), ("O_final", 1), ("T", 1), ("Q", 0)]
M_CBAM = {"1A": 100, "2A": 44, "2B": 44, "3A": 100, "3B": 100, "3C": 100}


def rnd(x, dp):
    q = 10 ** dp
    return (int(abs(x) * q + 0.5 + 1e-9) / q) * (1 if x >= 0 else -1)


def snapshot(wb):
    snap = {}
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        v = ur.Value
        snap[ws.Name] = (ur.Row, ur.Column, v if isinstance(v, tuple) else ((v,),))
    return snap


def max_diff(before, after, skip):
    worst, bad = 0.0, []
    for name, (r0, c0, vb) in before.items():
        ra, ca, va = after[name]
        for i, rowb in enumerate(vb):
            for j, b in enumerate(rowb):
                r, c = r0 + i, c0 + j
                if skip(name, r, c):
                    continue
                ia, ja = r - ra, c - ca
                a = va[ia][ja] if 0 <= ia < len(va) and 0 <= ja < len(va[0]) else None
                if isinstance(b, (int, float)) and isinstance(a, (int, float)):
                    d = abs(a - b)
                    worst = max(worst, d)
                    if d > 0 and len(bad) < 10:
                        bad.append((name, r, c, b, a))
                elif a != b:
                    worst = max(worst, 1.0)
                    if len(bad) < 10:
                        bad.append((name, r, c, b, a))
    return worst, bad


def set_elasticities(wb):
    mi = wb.Worksheets(MI)
    for k in range(8):
        r = 66 + k
        assert mi.Range("D%d" % r).Value == PRODUCTS[k], "row %d is %r" % (r, mi.Range("D%d" % r).Value)
        assert abs(mi.Range("E%d" % r).Value - (-0.5)) < 1e-12, "E%d is not -0.5" % r
        mi.Range("E%d" % r).Value = EPS_NEW[k]
        mi.Range("F%d" % r).Value = BASIS[k] + " Evidence note: egypt/supporting/OutputElasticity_Note_v0.1."
        mi.Range("G%d" % r).Value = "low-medium"
    mi.Range("D64").Value = (str(mi.Range("D64").Value) + " eps by product from v1.6 (OutputElasticity_Note_v0.1); "
                             "cement -0.10, steel and fertilisers -0.40, aluminium -0.50.")


def export_live(co, t2f, i):
    """One scenario's live Table 2 values (active scenario already set)."""
    v = lambda a: co.Range(a).Value
    return dict(J=v("D66"), K=v("D67"), L=v("D68"), M=v("D69"), N=v("D70"), P=v("D73"), Q=v("D74"), T=v("D75"),
                Onp=v("D72"), O_full=v("D71"), O_final=t2f.Range("D63").Value,
                dghg=v("D54"), cpat_blk=v("D48"), dB=v("D46") + v("D47"), D_nb=v("D96"), B=v("D44") + v("D45"),
                r=100 * v("D30"), i_f=100 * v("D31"))


def refresh_snapshot(wb, xl, mirror):
    wsS, co, t2f = wb.Worksheets("Settings"), wb.Worksheets(CO), wb.Worksheets(T2F)
    active0 = wsS.Range("B10").Value
    live = {}
    tol = dict(K=0.35, P=0.12, Q=45.0, T=0.7, N=0.08, J=0.02)
    warn = []
    for i, b in enumerate(BUNDLES):
        wsS.Range("B10").Value = b
        xl.CalculateFull()
        assert co.Range("D6").Value == b
        d = export_live(co, t2f, i)
        live[b] = d
        row = 80 + i
        for c, k in zip("CDEFGHIJKL", ["J", "K", "L", "M", "N", "O_full", "Onp", "P", "Q", "T"]):
            co.Range("%s%d" % (c, row)).Value = d[k]
        co.Range("M%d" % row).Value = round(d["O_final"], 2)
        for k, t in tol.items():
            if abs(d[k] - mirror[b][k]) > t:
                warn.append("%s %s: live %.3f, mirror %.3f" % (b, k, d[k], mirror[b][k]))
        print(b, " ".join("%s %.3f" % (k, d[k]) for k in ("K", "P", "Q", "N", "T", "O_final")))
    wsS.Range("B10").Value = active0
    xl.CalculateFull()
    return live, warn


def write_printed(wb, live):
    ws = wb.Worksheets(T2F)
    scen_cols = "DEFGHI"
    for i, (key, dp) in enumerate(METRICS):
        r = 19 + i
        for j, b in enumerate(BUNDLES):
            src = M_CBAM[b] if key == "M" else live[b][key]
            ws.Range("%s%d" % (scen_cols[j], r)).Value = rnd(src, dp)
    ws.Range("B17").Value = ("B. As printed in the final documents (typed): EGYPT_CarveOut_Table2_%s, results text %s Table 2" % (VER, VER))
    ws.Range("B1").Value = "Table 2 (2030), final %s: confirmation of the final numbers in one workbook" % VER
    co = wb.Worksheets(CO)
    t = str(co.Range("B1").Value)
    co.Range("B1").Value = t.replace("final v1.5", "final " + VER)
    co.Range("B78").Value = ("F. Stored 2030 snapshot, all bundles (refreshed in %s from the live sheet for all six; "
                             "= EGYPT_CarveOut_Table2_%s)" % (VER, VER))


def add_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "%s: product-specific output elasticities (Manual inputs E66:E73)" % VER
    ws.Range("A%d" % r0).Font.Bold = True
    items = [
        ("Table2_Final: figures printed in the final documents differ from the workbook (count; expected 0)",
         "=%s!D42" % T2F),
        ("Output elasticities: cement E69 (-0.10 expected)", "='%s'!E69" % MI),
        ("Output elasticities: sum of the other seven (-2.9 expected: 3 x -0.4 + 3 x -0.4 + -0.5)",
         "=SUM('%s'!E66:E68)+SUM('%s'!E70:E73)" % (MI, MI)),
    ]
    for i, (lab, f) in enumerate(items):
        ws.Range("A%d" % (r0 + 1 + i)).Value = lab
        ws.Range("D%d" % (r0 + 1 + i)).Formula = f
        ws.Range("D%d" % (r0 + 1 + i)).Interior.Color = REVIEW
    return r0 + 1


def settings(wb, chk, warn):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt, final " + VER
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Output elasticity by product (Manual inputs E66:E73): cement -0.10, steel and fertilisers -0.40, aluminium -0.50 "
        "(was -0.5 for all; OutputElasticity_Note_v0.1). Fund fixed point re-solved, Task J v2 table, Table2_Industry and "
        "CarveOut_Table2 stored snapshots refilled. Every scenario-2 result moves; LEGACY/baseline unchanged. Check row %d.%s"
        % (chk, " Mirror difference warnings: %d (documents to be regenerated)." % len(warn) if warn else ""))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    assert os.path.exists(SRC), "kernel v1.5 is missing: " + SRC
    mirror = json.load(open(os.path.join(ADHOC, "carveout_v1_6_results.json"), encoding="utf8"))["results"]
    if os.path.exists(DST):
        os.remove(DST)
    if not os.path.exists(SRC_OLD):
        shutil.copyfile(SRC, SRC_OLD)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    ok = False
    try:
        wb = xl.Workbooks.Open(DST)
        xl.CalculateFull()
        assert wb.Worksheets("Settings").Range("B10").Value == "LEGACY", "set Settings!B10 to LEGACY in v1.5 first"
        before = snapshot(wb)
        chk_last_before = before["Check"][0] + len(before["Check"][2]) - 1

        set_elasticities(wb)
        xl.CalculateFull()
        print("v2 rows refilled:", b15.refill_v2(wb))
        print("fund fixed point:", b16.solve_fund(wb))
        k = b16.t2_rows(wb)
        b16.refill_table2(wb, k)
        xl.CalculateFull()
        live, warn = refresh_snapshot(wb, xl, mirror)
        write_printed(wb, live)
        chk = add_check(wb)
        xl.CalculateFull()
        after = snapshot(wb)

        # LEGACY-state regression: only the intended areas may differ
        def skip(name, r, c):
            if name == "Settings":
                return True
            if name == MI and r in range(64, 74):
                return True
            if name == "Check" and r > chk_last_before:
                return True
            if name in (CO, T2F):
                return True                                  # stored snapshots / printed figures: intended
            if name in ("Table2_Industry", "Fund_Industry"):
                return True                                  # stored snapshot rows and stored fund rows: intended
            if name == "Check":
                return True                                  # Task J v2 stored table: intended
            return False
        diff, bad = max_diff(before, after, skip)
        print("max abs diff in the other sheets (LEGACY state): %.6g" % diff)
        for x in bad:
            print("  DIFF", x)
        mism = wb.Worksheets(T2F).Range("D42").Value
        print("printed-vs-workbook mismatches:", mism)
        json.dump(live, open(os.path.join(ADHOC, "carveout_v1_6_results_live.json"), "w"), indent=1)
        if warn:
            print("MIRROR DIFFERENCES (live vs make_carveout_v0_5.py):")
            for w in warn:
                print("  ", w)
        if diff != 0:
            print("NOT saved: unexpected change outside the intended areas.")
            wb.Close(False)
            sys.exit(1)
        if mism != 0 or warn:
            print("DOCUMENTS MUST BE REGENERATED from carveout_v1_6_results_live.json (printed figures differ from the workbook).")
        settings(wb, chk, warn)
        wb.Save()
        wb.Close(False)
        ok = True
    finally:
        xl.Quit()
    if ok:
        os.replace(SRC, SRC_OLD)
        os.makedirs(FINAL, exist_ok=True)
        old_final = os.path.join(FINAL, "CPAT_Industry_Kernel_Egypt_v1.5.xlsx")
        if os.path.exists(old_final):
            shutil.move(old_final, os.path.join(ARCHIVE, "CPAT_Industry_Kernel_Egypt_v1.5.xlsx"))
        shutil.copyfile(DST, os.path.join(FINAL, os.path.basename(DST)))
        ef = os.path.join(REPO, "egypt-final", "Egypt_Model.xlsx")
        if os.path.isdir(os.path.dirname(ef)):
            shutil.copyfile(DST, ef)
        print("saved", DST, "; copied to egypt/final and egypt-final/Egypt_Model.xlsx; v1.5 moved to Old/ and egypt/archive")


if __name__ == "__main__":
    main()
